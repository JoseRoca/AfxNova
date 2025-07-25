' ########################################################################################
' Microsoft Windows
' File: GLCTX_Indexed Geometry.bas
' Contents: Graphic Control OpenGL - Indexed geometry
' Compiler: FreeBasic 32 & 64 bit
' Demonstrates how to optimize performance by using indexed geometry. As a demonstration,
' the sample reduces the vertex count of a simple cube from 24 to 8 by redefining the
' cube�s geometry using an indices array.
' Based on ogl_indexed_geometry.cpp, by Kevin Harris, 02/01/05.
' Adapted in 2017 by Jos� Roca. Freeware. Use at your own risk.
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

CONST GL_WINDOWWIDTH   = 600                          ' Window width
CONST GL_WINDOWHEIGHT  = 400                          ' Window height
CONST GL_WindowCaption = "OpenGL: Indexes geometry"   ' Window caption
CONST IDC_GRCTX = 1001

' Vertex structure
TYPE Vertex
   r AS SINGLE
   g AS SINGLE
   b AS SINGLE
   x AS SINGLE
   y AS SINGLE
   z AS SINGLE
END TYPE

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

' =======================================================================================
' OpenGL class
' =======================================================================================
TYPE CtxOgl

   Private:
      m_pGraphCtx AS CGraphCtx PTR

      m_bUseIndexedGeometry AS BOOLEAN = TRUE
      m_fSpinX AS SINGLE
      m_fSpinY AS SINGLE

      '// To understand how indexed geometry works, we must first build something
      '// which can be optimized through the use of indices.
      '//
      '// Below, is the vertex data for a simple multi-colored cube, which is defined
      '// as 6 individual quads, one quad for each of the cube's six sides. At first,
      '// this doesn�t seem too wasteful, but trust me it is.
      '//
      '// You see, we really only need 8 vertices to define a simple cube, but since
      '// we're using a quad list, we actually have to repeat the usage of our 8
      '// vertices 3 times each. To make this more understandable, I've actually
      '// numbered the vertices below so you can see how the vertices get repeated
      '// during the cube's definition.
      '//
      '// Note how the first 8 vertices are unique. Everting else after that is just
      '// a repeat of the first 8.
      DIM m_cubeVertices(23) AS Vertex = { _
         (1.0,0.0,0.0, -1.0,-1.0, 1.0), _  ' // 0 (unique)
         (0.0,1.0,0.0,  1.0,-1.0, 1.0), _  ' // 1 (unique)
         (0.0,0.0,1.0,  1.0, 1.0, 1.0), _  ' // 2 (unique)
         (1.0,1.0,0.0, -1.0, 1.0, 1.0), _  ' // 3 (unique)
         (1.0,0.0,1.0, -1.0,-1.0,-1.0), _  ' // 4 (unique)
         (0.0,1.0,1.0, -1.0, 1.0,-1.0), _  ' // 5 (unique)
         (1.0,1.0,1.0,  1.0, 1.0,-1.0), _  ' // 6 (unique)
         (1.0,0.0,0.0,  1.0,-1.0,-1.0), _  ' // 7 (unique)
         (0.0,1.0,1.0, -1.0, 1.0,-1.0), _  ' // 5 (start repeating here)
         (1.0,1.0,0.0, -1.0, 1.0, 1.0), _  ' // 3 (repeat of vertex 3)
         (0.0,0.0,1.0,  1.0, 1.0, 1.0), _  ' // 2 (repeat of vertex 2... etc.)
         (1.0,1.0,1.0,  1.0, 1.0,-1.0), _  ' // 6
         (1.0,0.0,1.0, -1.0,-1.0,-1.0), _  ' // 4
         (1.0,0.0,0.0,  1.0,-1.0,-1.0), _  ' // 7
         (0.0,1.0,0.0,  1.0,-1.0, 1.0), _  ' // 1
         (1.0,0.0,0.0, -1.0,-1.0, 1.0), _  ' // 0
         (1.0,0.0,0.0,  1.0,-1.0,-1.0), _  ' // 7
         (1.0,1.0,1.0,  1.0, 1.0,-1.0), _  ' // 6
         (0.0,0.0,1.0,  1.0, 1.0, 1.0), _  ' // 2
         (0.0,1.0,0.0,  1.0,-1.0, 1.0), _  ' // 1
         (1.0,0.0,1.0, -1.0,-1.0,-1.0), _  ' // 4
         (1.0,0.0,0.0, -1.0,-1.0, 1.0), _  ' // 0
         (1.0,1.0,0.0, -1.0, 1.0, 1.0), _  ' // 3
         (0.0,1.0,1.0, -1.0, 1.0,-1.0)}    ' // 5

      '// Now, to save ourselves the bandwidth of passing a bunch or redundant vertices
      '// down the graphics pipeline, we shorten our vertex list and pass only the
      '// unique vertices. We then create a indices array, which contains index values
      '// that reference vertices in our vertex array.
      '//
      '// In other words, the vertex array doens't actually define our cube anymore,
      '// it only holds the unique vertices; it's the indices array that now defines
      '// the cube's geometry.
      DIM m_cubeVertices_indexed(7) AS Vertex = { _
         (1.0,0.0,0.0,  -1.0,-1.0, 1.0), _  ' // 0
         (0.0,1.0,0.0,   1.0,-1.0, 1.0), _  ' // 1
         (0.0,0.0,1.0,   1.0, 1.0, 1.0), _  ' // 2
         (1.0,1.0,0.0,  -1.0, 1.0, 1.0), _  ' // 3
         (1.0,0.0,1.0,  -1.0,-1.0,-1.0), _  ' // 4
         (0.0,1.0,1.0,  -1.0, 1.0,-1.0), _  ' // 5
         (1.0,1.0,1.0,   1.0, 1.0,-1.0), _  ' // 6
         (1.0,0.0,0.0,   1.0,-1.0,-1.0)}    ' // 7

      DIM m_cubeIndices(23) AS DWORD = { 0, 1, 2, 3, 4, 5, 6, 7, 5, 3, 2, 6, 4, 7, 1, 0, 7, 6, 2, 1, 4, 0, 3, 5 }

      '// Note: While the cube above makes for a good example of how indexed geometry
      '//       works. There are many situations which can prevent you from using
      '//       an indices array to its full potential.
      '//
      '//       For example, if our cube required normals for lighting, things would
      '//       become problematic since each vertex would be shared between three
      '//       faces of the cube. This would not give you the lighting effect that
      '//       you really want since the best you could do would be to average the
      '//       normal's value between the three faces which used it.
      '//
      '//       Another example would be texture coordinates. If our cube required
      '//       unique texture coordinates for each face, you really wouldn�t gain
      '//       much from using an indices array since each vertex would require a
      '//       different texture coordinate depending on which face it was being
      '//       used in.

   Public:
      DECLARE CONSTRUCTOR (BYVAL pGraphCtx AS CGraphCtx PTR)
      DECLARE DESTRUCTOR
      DECLARE SUB SetupScene
      DECLARE SUB ResizeScene
      DECLARE SUB RenderScene
      DECLARE FUNCTION ProcessKeystrokes (BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS BOOLEAN
      DECLARE FUNCTION ProcessMouse (BYVAL uMsg AS UINT, BYVAL wKeyState AS DWORD, BYVAL x AS LONG, BYVAL y AS LONG) AS BOOLEAN

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

   ' Specify clear values for the color buffers
   glClearColor 0.0, 0.0, 0.0, 0.0
   ' Specify the clear value for the depth buffer
   glClearDepth 1.0
   ' Specify the value used for depth-buffer comparisons
   glDepthFunc GL_LESS
   ' Enable depth comparisons and update the depth buffer
   glEnable GL_DEPTH_TEST
   ' Select smooth shading
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

   glTranslatef(0.0!, 0.0!, -5.0!)
   glRotatef(-m_fSpinY, 1.0!, 0.0!, 0.0!)
   glRotatef(-m_fSpinX, 0.0!, 0.0!, 1.0!)

   IF m_bUseIndexedGeometry THEN
      glInterleavedArrays(GL_C3F_V3F, 0, @m_cubeVertices_indexed(0))
      ' Note: The original C program uses GL_UNSIGNED_BYTE, but PB only supports
      ' signed bytes, so we are using an array of DWORDs
      glDrawElements(GL_QUADS, 24, GL_UNSIGNED_INT, @m_cubeIndices(0))
   ELSE
      glInterleavedArrays(GL_C3F_V3F, 0, @m_cubeVertices(0))
      glDrawArrays(GL_QUADS, 0, 24)
   END IF

   ' // Required: force execution of GL commands in finite time
   glFlush

   ' // Required: Force repainting of the control
   IF m_pGraphCtx THEN InvalidateRect(m_pGraphCtx->hWindow, NULL, CTRUE)

END SUB
' =======================================================================================

' =======================================================================================
' Processes keystrokes
' Parameters:
' * uMsg = The Windows message
' * wParam = Additional message information.
' * lParam = Additional message information.
' The contents of the wParam and lParam parameters depend on the value of the uMsg parameter. 
' Return value: TRUE if the message has been processed, or FALSE otherwise.
' =======================================================================================
FUNCTION CtxOgl.ProcessKeystrokes (BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS BOOLEAN

   STATIC bRenderInWireFrame AS BOOLEAN

   FUNCTION = FALSE
   SELECT CASE uMsg
      CASE WM_KEYDOWN
         SELECT CASE LOWORD(wParam)
            CASE VK_F1
               m_bUseIndexedGeometry = NOT m_bUseIndexedGeometry
               FUNCTION = TRUE
         END SELECT
   END SELECT

END FUNCTION
' =======================================================================================

' =======================================================================================
' Processes mouse clicks and movement
' Parameters:
' * wMsg      = Windows message
' * wKeyState = Indicates whether various virtual keys are down.
'               MK_CONTROL    The CTRL key is down.
'               MK_LBUTTON    The left mouse button is down.
'               MK_MBUTTON    The middle mouse button is down.
'               MK_RBUTTON    The right mouse button is down.
'               MK_SHIFT      The SHIFT key is down.
'               MK_XBUTTON1   Windows 2000/XP: The first X button is down.
'               MK_XBUTTON2   Windows 2000/XP: The second X button is down.
' * x         = x-coordinate of the cursor
' * y         = y-coordinate of the cursor
' Return value: TRUE if the message has been processed, or FALSE otherwise.
' =======================================================================================
FUNCTION CtxOgl.ProcessMouse (BYVAL uMsg AS UINT, BYVAL wKeyState AS DWORD, BYVAL x AS LONG, BYVAL y AS LONG) AS BOOLEAN

   STATIC ptLastMousePosit AS POINT
   STATIC ptCurrentMousePosit AS POINT
   STATIC bMousing AS BOOLEAN

   SELECT CASE uMsg

      CASE WM_LBUTTONDOWN
         ptLastMousePosit.x = x
         ptCurrentMousePosit.x = x
         ptLastMousePosit.y = y
         ptCurrentMousePosit.y = y
         bMousing = TRUE
         RETURN TRUE

      CASE WM_LBUTTONUP
         bMousing = FALSE
         RETURN TRUE

      CASE WM_MOUSEMOVE
         ptCurrentMousePosit.x = x
         ptCurrentMousePosit.y = y
         IF bMousing THEN
            m_fSpinX -= (ptCurrentMousePosit.x - ptLastMousePosit.x)
            m_fSpinY -= (ptCurrentMousePosit.y - ptLastMousePosit.y)
         END IF
         ptLastMousePosit.x = ptCurrentMousePosit.x
         ptLastMousePosit.y = ptCurrentMousePosit.y
         RETURN TRUE

   END SELECT

END FUNCTION
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

      CASE WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE
         DIM pCtxOgl AS CtxOgl PTR = cast(CtxOgl PTR, dwRefData)
         IF pCtxOgl THEN pCtxOgl->ProcessMouse(uMsg, wParam, LOWORD(lParam), HIWORD(lParam))
         EXIT FUNCTION

      CASE WM_KEYDOWN
         SELECT CASE GET_WM_COMMAND_ID(wParam, lParam)
            CASE VK_ESCAPE
               SendMessageW(GetParent(hwnd), WM_CLOSE, 0, 0)
               EXIT FUNCTION
            CASE ELSE
               DIM pCtxOgl AS CtxOgl PTR = cast(CtxOgl PTR, dwRefData)
               IF pCtxOgl->ProcessKeystrokes(uMsg, wParam, lParam) = TRUE THEN EXIT FUNCTION
         END SELECT

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
