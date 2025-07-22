' ########################################################################################
' Microsoft Windows
' Contents: CWindow OpenGL Graphic Control Skeleton
' Graphic control, GDI+ classes and OpenGL working together
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
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
' The following sample code draws a line.
' Change it with your own code.
' ========================================================================================
SUB GDIP_Render (BYVAL hDC AS HDC)

   ' // Create a graphics object from the window device context
   DIM graphics AS CGpGraphics = hdc
   ' // Get the DPI scaling ratios
   DIM rxRatio AS SINGLE = graphics.GetDpiX / 96
   DIM ryRatio AS SINGLE = graphics.GetDpiY / 96
   ' // Set the scale transform
   graphics.ScaleTransform(rxRatio, ryRatio)

   ' // Create a Pen object
   DIM pen AS CGpPen = CGpPen(GDIP_ARGB(255, 0, 255, 0), 15)

   ' // Draw a line
   graphics.DrawLine(@pen, 0, 0, 200, 100)

END SUB
' ========================================================================================

' =======================================================================================
' All the setup goes here
' =======================================================================================
SUB RenderScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)

   ' // Specify clear values for the color buffers
   glClearColor 0.0, 0.0, 0.0, 0.0
   ' // Specify the clear value for the depth buffer
   glClearDepth 1.0
   ' // Specify the value used for depth-buffer comparisons
   glDepthFunc GL_LESS
   ' // Enable depth comparisons and update the depth buffer
   glEnable GL_DEPTH_TEST
   ' // Select smooth shading
   glShadeModel GL_SMOOTH

   ' // Prevent divide by zero making height equal one
   IF nHeight = 0 THEN nHeight = 1
   ' // Reset the current viewport
   glViewport 0, 0, nWidth, nHeight
   ' // Select the projection matrix
   glMatrixMode GL_PROJECTION
   ' // Reset the projection matrix
   glLoadIdentity
   ' // Calculate the aspect ratio of the window
   gluPerspective 45.0!, nWidth / nHeight, 0.1!, 100.0!
   ' // Select the model view matrix
   glMatrixMode GL_MODELVIEW
   ' // Reset the model view matrix
   glLoadIdentity

   ' // Clear the screen buffer
   glClear GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT
   ' // Reset the view
   glLoadIdentity

   glTranslatef -1.5!, 0.0!, -6.0!       ' Move left 1.5 units and into the screen 6.0

   glBegin GL_TRIANGLES                 ' Drawing using triangles
      glColor3f   1.0!, 0.0!, 0.0!       ' Set the color to red
      glVertex3f  0.0!, 1.0!, 0.0!       ' Top
      glColor3f   0.0!, 1.0!, 0.0!       ' Set the color to green
      glVertex3f  1.0!,-1.0!, 0.0!       ' Bottom right
      glColor3f   0.0!, 0.0!, 1.0!       ' Set the color to blue
      glVertex3f -1.0!,-1.0!, 0.0!       ' Bottom left
   glEnd                                 ' Finished drawing the triangle

   glTranslatef 3.0!,0.0!,0.0!           ' Move right 3 units

   glColor3f 0.5!, 0.5!, 1.0!            ' Set the color to blue one time only
   glBegin GL_QUADS                      ' Draw a quad
      glVertex3f -1.0!, 1.0!, 0.0!       ' Top left
      glVertex3f  1.0!, 1.0!, 0.0!       ' Top right
      glVertex3f  1.0!,-1.0!, 0.0!       ' Bottom right
      glVertex3f -1.0!,-1.0!, 0.0!       ' Bottom left
   glEnd                                 ' Done drawing the quad

   ' // Required: force execution of GL commands in finite time
   glFlush

END SUB
' =======================================================================================

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
   DIM pWindow AS CWindow
   pWindow.Create(NULL, "CWindow+CGraphCtx+CGdiPlus+OpenGL", @WndProc)
   pWindow.SetClientSize(400, 250)
   pWindow.Center

   ' // Add a graphic control with OPENGL enabled
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "OPENGL", 0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
   ' // Anchor the control
   pWindow.AnchorControl(IDC_GRCTX, AFX_ANCHOR_HEIGHT_WIDTH)

   ' // Clear the background
   pGraphCtx.Clear RGB_WHITE
   ' // Make the control stretchable
   pGraphCtx.Stretchable = TRUE
   ' // Make current the rendering context
   pGraphCtx.MakeCurrent
   ' // Render the OpenGL scene
   RenderScene pGraphCtx.GetVirtualBufferWidth, pGraphCtx.GetVirtualBufferHeight
   ' // Draw the GDI+ graphics
   GDIP_Render pGraphCtx.GetMemDC

   ' // Dispatch Windows events
   FUNCTION = pWindow.DoEvents(nCmdShow)

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window callback procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   SELECT CASE uMsg

      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

    	CASE WM_DESTROY
         ' // End the application
         PostQuitMessage(0)
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hWnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================