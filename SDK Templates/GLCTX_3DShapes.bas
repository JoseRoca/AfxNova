' ########################################################################################
' Microsoft Windows
' File: GLCTX_3DShapes.bas
' Contents: Graphic Control OpenGL: 3D shapes
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 Jos� Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
#INCLUDE ONCE "AfxNova/CGraphCtx.inc"
USING AfxNova

DECLARE FUNCTION wWinMain (BYVAL hInstance AS HINSTANCE, _
                           BYVAL hPrevInstance AS HINSTANCE, _
                           BYVAL pwszCmdLine AS WSTRING PTR, _
                           BYVAL nCmdShow AS LONG) AS LONG
   END wWinMain(GetModuleHandleW(NULL), NULL, wCommand(), SW_NORMAL)

' // Forward declarations
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT
DECLARE FUNCTION GraphCtx_SubclassProc ( _
   BYVAL hwnd   AS HWND, _                 ' // Control window handle
   BYVAL uMsg   AS UINT, _                 ' // Type of message
   BYVAL wParam AS WPARAM, _               ' // First message parameter
   BYVAL lParam AS LPARAM, _               ' // Second message parameter
   BYVAL uIdSubclass AS UINT_PTR, _        ' // The subclass ID
   BYVAL dwRefData AS DWORD_PTR _          ' // Pointer to reference data
   ) AS LRESULT

CONST GL_WINDOWWIDTH   = 600                    ' Window width
CONST GL_WINDOWHEIGHT  = 400                    ' Window height
CONST GL_WindowCaption = "OpenGL - 3D shapes"   ' Window caption
CONST IDC_GRCTX = 1001

' =======================================================================================
' OpenGL class
' =======================================================================================
TYPE CtxOgl

   Private:
      m_pGraphCtx AS CGraphCtx PTR
      rtri AS SINGLE
      rquad AS SINGLE

   Public:
      DECLARE CONSTRUCTOR (BYVAL pGraphCtx AS CGraphCtx PTR)
      DECLARE DESTRUCTOR
      DECLARE SUB SetupScene
      DECLARE SUB ResizeScene
      DECLARE SUB RenderScene

END TYPE
' =======================================================================================

' ========================================================================================
' COGL constructor
' ========================================================================================
CONSTRUCTOR CtxOgl (BYVAL pGraphCtx AS CGraphCtx PTR)
   m_pGraphCtx = pGraphCtx
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' COGL Destructor
' ========================================================================================
DESTRUCTOR CtxOgl
END DESTRUCTOR
' ========================================================================================

' =======================================================================================
' All the setup goes here
' =======================================================================================
SUB CtxOgl.SetupScene

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

END SUB
' =======================================================================================

' =======================================================================================
SUB CtxOgl.ResizeScene

   ' // Get the dimensions of the control
   IF m_pGraphCtx = NULL THEN EXIT SUB
   DIM nWidth AS LONG = AfxGetWindowWidth(m_pGraphCtx->hWindow)
   DIM nHeight AS LONG = AfxGetWindowHeight(m_pGraphCtx->hWindow)
   ' // Prevent divide by zero making height equal one
   IF nHeight = 0 THEN nHeight = 1
   ' // Reset the current viewport
   glViewport 0, 0, nWidth, nHeight
   ' // Select the projection matrix
   glMatrixMode GL_PROJECTION
   ' // Reset the projection matrix
   glLoadIdentity
   ' // Calculate the aspect ratio of the window
   gluPerspective 45.0, nWidth / nHeight, 0.1, 100.0
   ' // Select the model view matrix
   glMatrixMode GL_MODELVIEW
   ' // Reset the model view matrix
   glLoadIdentity

END SUB
' =======================================================================================

' =======================================================================================
SUB CtxOgl.RenderScene

   ' // Clear the screen buffer
   glClear GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT
   ' // Reset the view
   glLoadIdentity

   glTranslatef -1.5, 0.0, -6.0           ' Move left 1.5 units and into the screen
   glRotatef rtri, 0.0, 1.0, 0.0          ' Rotate the triangle on the Y axis

   glBegin GL_TRIANGLES
      ' Front
      glColor3f   1.0,  0.0,  0.0         ' Red
      glVertex3f  0.0,  1.0,  0.0         ' Top of triangle (Front)
      glColor3f   0.0,  1.0,  0.0         ' Green
      glVertex3f -1.0, -1.0,  1.0         ' Left of triangle (Front)
      glColor3f   0.0,  0.0,  1.0         ' Blue
      glVertex3f  1.0, -1.0,  1.0         ' Right of triangle (Front)

      ' Right
      glColor3f   1.0,  0.0,  0.0         ' Red
      glVertex3f  0.0,  1.0,  0.0         ' Top of triangle (Right)
      glColor3f   0.0,  0.0,  1.0         ' Blue
      glVertex3f  1.0, -1.0,  1.0         ' Left of triangle (Right)
      glColor3f   0.0,  1.0,  0.0         ' Green
      glVertex3f  1.0, -1.0, -1.0         ' Right of triangle (Right)

      ' Back
      glColor3f   1.0,  0.0,  0.0         ' Red
      glVertex3f  0.0,  1.0,  0.0         ' Top of triangle (Back)
      glColor3f   0.0,  1.0,  0.0         ' Green
      glVertex3f  1.0, -1.0, -1.0         ' Left of triangle (Back)
      glColor3f   0.0,  0.0,  1.0         ' Blue
      glVertex3f -1.0, -1.0, -1.0         ' Right of triangle (Back)

      ' Left
      glColor3f   1.0,  0.0,  0.0         ' Red
      glVertex3f  0.0,  1.0,  0.0         ' Top of triangle (Left)
      glColor3f   0.0,  0.0,  1.0         ' Blue
      glVertex3f -1.0, -1.0, -1.0         ' Left of triangle (Left)
      glColor3f   0.0,  1.0,  0.0         ' Green
      glVertex3f -1.0, -1.0,  1.0         ' Right of triangle (Left)
   glEnd

   glLoadIdentity
   glTranslatef 1.5, 0.0, -7.0            ' Move right 1.5 units and into the screen
   glRotatef rquad, 1.0, 1.0, 1.0         ' Rotate the quad on the X axis

   glBegin GL_QUADS
      glColor3f   0.0,  1.0,  0.0         ' Set the color to green
      glVertex3f  1.0,  1.0, -1.0         ' Top right of the quad (Top)
      glVertex3f -1.0,  1.0, -1.0         ' Top left of the quad (Top)
      glVertex3f -1.0,  1.0,  1.0         ' Bottom left of the quad (Top)
      glVertex3f  1.0,  1.0,  1.0         ' Bottom right of the quad (Top)

      glColor3f   1.0,  0.5,  0.0         ' Set the color to orange
      glVertex3f  1.0, -1.0,  1.0         ' Top right of the quad (Bottom)
      glVertex3f -1.0, -1.0,  1.0         ' Top left of the quad (Bottom)
      glVertex3f -1.0, -1.0, -1.0         ' Bottom left of the quad (Bottom)
      glVertex3f  1.0, -1.0, -1.0         ' Bottom right of the quad (Bottom)

      glColor3f   1.0,  0.0,  0.0         ' Set the color to red
      glVertex3f  1.0,  1.0,  1.0         ' Top right of the quad (Front)
      glVertex3f -1.0,  1.0,  1.0         ' Top left of the quad (Front)
      glVertex3f -1.0, -1.0,  1.0         ' Bottom left of the quad (Front)
      glVertex3f  1.0, -1.0,  1.0         ' Bottom right of the quad (Front)

      glColor3f   1.0,  1.0,  0.0         ' Set the color to yellow
      glVertex3f  1.0, -1.0, -1.0         ' Top right of the quad (Back)
      glVertex3f -1.0, -1.0, -1.0         ' Top left of the quad (Back)
      glVertex3f -1.0,  1.0, -1.0         ' Bottom left of the quad (Back)
      glVertex3f  1.0,  1.0, -1.0         ' Bottom right of the quad (Back)

      glColor3f   0.0,  0.0,  1.0         ' Set the color to blue
      glVertex3f -1.0,  1.0,  1.0         ' Top right of the quad (Left)
      glVertex3f -1.0,  1.0, -1.0         ' Top left of the quad (Left)
      glVertex3f -1.0, -1.0, -1.0         ' Bottom left of the quad (Left)
      glVertex3f -1.0, -1.0,  1.0         ' Bottom right of the quad (Left)

      glColor3f   1.0,  0.0,  1.0         ' Set the color to violet
      glVertex3f  1.0,  1.0, -1.0         ' Top right of the quad (Right)
      glVertex3f  1.0,  1.0,  1.0         ' Top left of the quad (Right)
      glVertex3f  1.0, -1.0,  1.0         ' Bottom left of the quad (Right)
      glVertex3f  1.0, -1.0, -1.0         ' Bottom right of the quad (Right)
   glEnd

   rtri = rtri + 0.2                      ' Increase the rotation variable for the triangle
   rquad = rquad - 0.15                   ' Decrease the rotation variable for the quad

   ' // Required: force execution of GL commands in finite time
   glFlush

   ' // Required: Force repainting of the control
   IF m_pGraphCtx THEN InvalidateRect(m_pGraphCtx->hWindow, NULL, CTRUE)

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

   ' // Creates the main window
   DIM pWindow AS CWindow = "MyClassName"   ' Use the name you wish
   ' // Create the window
   DIM hwndMain AS HWND = pWindow.Create(NULL, GL_WindowCaption, @WndProc)
   ' // Don't erase the background
   ' // Use a black brush
   pWindow.Brush = CreateSolidBrush(BGR(255, 255, 255))
   ' // Sizes the window by setting the wanted width and height of its client area
   pWindow.SetClientSize(GL_WINDOWWIDTH, GL_WINDOWHEIGHT)
   ' // Centers the window
   pWindow.Center

   ' // Add a subclassed graphic control with OPENGL enabled
   DIM pGraphCtx AS CGraphCtx = CGraphCtx(@pWindow, IDC_GRCTX, "OPENGL", _
       0, 0, pWindow.ClientWidth, pWindow.ClientHeight)
'   pGraphCtx.Stretchable = TRUE
   pGraphCtx.Resizable = TRUE
   ' // Set the focus in the control
   SetFocus pGraphCtx.hWindow
   ' // Set the timer (using a timer to trigger redrawing allows a smoother rendering)
   SetTimer(pGraphCtx.hWindow, 1, 0, NULL)
   ' // Anchor the control
   pWindow.AnchorControl(IDC_GRCTX, AFX_ANCHOR_HEIGHT_WIDTH)

   ' // Create an instance of the CtxOgl class
   DIM pCtxOgl AS CtxOgl = @pGraphCtx
   ' // Subclass the graphic control
   SetWindowSubclass pGraphCtx.hWindow, CAST(SUBCLASSPROC, @GraphCtx_SubclassProc), IDC_GRCTX, CAST(DWORD_PTR, @pCtxOgl)
   ' // Setup the OpenGL scene
   pCtxOgl.SetupScene
   ' // Resize the OpenGL scene
   pCtxOgl.ResizeScene
   ' // Render the OpenGL scene
   pCtxOgl.RenderScene

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
         RETURN 0

      CASE WM_SYSCOMMAND
         ' // Disable the Windows screensaver
         IF (wParam AND &hFFF0) = SC_SCREENSAVE THEN RETURN 0
         ' // Close the window
         IF (wParam AND &hFFF0) = SC_CLOSE THEN
            SendMessageW hwnd, WM_CLOSE, 0, 0
            RETURN 0
         END IF

      ' // Sent when the user selects a command item from a menu, when a control sends a
      ' // notification message to its parent window, or when an accelerator keystroke is translated.
      CASE WM_COMMAND
         SELECT CASE CBCTL(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF CBCTLMSG(wParam, lParam) = BN_CLICKED THEN SendMessageW(hwnd, WM_CLOSE, 0, 0)
         END SELECT
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

' ========================================================================================
' Processes messages for the subclassed graphic control
' ========================================================================================
FUNCTION GraphCtx_SubclassProc ( _
   BYVAL hwnd   AS HWND, _                 ' // Control window handle
   BYVAL uMsg   AS UINT, _                 ' // Type of message
   BYVAL wParam AS WPARAM, _               ' // First message parameter
   BYVAL lParam AS LPARAM, _               ' // Second message parameter
   BYVAL uIdSubclass AS UINT_PTR, _        ' // The subclass ID
   BYVAL dwRefData AS DWORD_PTR _          ' // Pointer to reference data
   ) AS LRESULT

   SELECT CASE uMsg

      CASE WM_GETDLGCODE
         ' // All keyboard input
         FUNCTION = DLGC_WANTALLKEYS
         EXIT FUNCTION

      CASE WM_COMMAND
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE IDCANCEL
               ' // If ESC key pressed, close the application by sending an WM_CLOSE message
               IF GET_WM_COMMAND_CMD(wParam, lParam) = BN_CLICKED THEN
                  SendMessageW hwnd, WM_CLOSE, 0, 0
                  EXIT FUNCTION
               END IF
         END SELECT

      CASE WM_LBUTTONDOWN
         MessageBoxW(GetParent(hwnd), "Click", "FreeBasic", MB_OK)
         EXIT FUNCTION

      CASE WM_TIMER
         ' // Render the scene
         DIM pCtxOgl AS CtxOgl PTR = cast(CtxOgl PTR, dwRefData)
         IF pCtxOgl THEN pCtxOgl->RenderScene
         EXIT FUNCTION

      CASE WM_SIZE
         ' // First perform the default action
         DefSubclassProc(hwnd, uMsg, wParam, lParam)
         ' // Check if the graphic contol is resizable
         DIM bResizable AS BOOLEAN =  AfxCGraphCtxPtr(hwnd)->Resizable
         ' // If it is resizable, we need to recreate the scene
         ' // because the rendering context has changed
         IF bResizable THEN
            DIM pCtxOgl AS CtxOgl PTR = cast(CtxOgl PTR, dwRefData)
            pCtxOgl->SetUpScene
            pCtxOgl->ResizeScene
            pCtxOgl->RenderScene
         END IF
         EXIT FUNCTION

      CASE WM_DESTROY
         ' // Kill the timer
         KillTimer(hwnd, 1)
         ' // REQUIRED: Remove control subclassing
         RemoveWindowSubclass hwnd, @GraphCtx_SubclassProc, uIdSubclass

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefSubclassProc(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
