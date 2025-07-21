# OpenGL Procedures

Assorted helper OpenGL procedures.

Freeglut 2.8.1 includes twenty two routines for generating easily-recognizable 3-d geometric objects. These routines are effectively the same ones that are included in the GLUT library, and reflect the functionality available in the aux toolkit described in the OpenGL Programmer's Guide. They are included to allow programmers to create with a single line of code a three-dimensional object which can be used to test a variety of OpenGL functionality.  None of the routines generates a display list for the object which it draws. The functions generate normals appropriate for lighting but, except for the teapot functions, do not generate texture coordinates. Do note that depth testing (GL_LESS) should be enabled for the correct drawing of the nonconvex objects, i.e., the glutTorus and glutTeapot.

I have adapted these routines to FreeBasic and grouped them in the include file AfxGlut.inc

To load textures for use with OpenGL, both from disk or a resource file, you can use the overloaded wrapper function **AfxGdipLoadTexture**, included in the file AfxGdiPlus.inc.

| Name       | Description |
| ---------- | ----------- |
| [AfxGlutSolidCone](#afxglutcone) | Draws a solid cone. |
| [AfxGlutWireCone](#afxglutcone) | Draws a wireframe cone. |
| [AfxGlutSolidCube](#afxglutcube) | Draws a solid cube. |
| [AfxGlutWireCube](#afxglutcube) | Draws a wireframe cube. |
| [AfxGlutSolidCylinder](#afxglutcylinder) | Draws a solid cylinder. |
| [AfxGlutWireCylinder](#afxglutcylinder) | Draws a wireframe cylinder. |
| [AfxGlutSolidDodecahedron](#afxglutdodecahedron) | Draws a solid dodecahedron. |
| [AfxGlutWireDodecahedron](#afxglutdodecahedron) | Draws a wireframe dodecahedron. |
| [AfxGlutSolidIcosahedron](#afxgluticosahedron) | Draws a solid icosahedron. |
| [AfxGlutWireIcosahedron](#afxgluticosahedron) | Draws a wireframe icosahedron. |
| [AfxGlutSolidOctahedron](#afxglutoctahedron) | Draws a solid octahedron. |
| [AfxGlutWireOctahedron](#afxglutoctahedron) | Draws a wireframe octahedron. |
| [AfxGlutSolidRhombicDodecahedron](#afxglutrhombicdodecahedron) | Draws a solid rhombic dodecahedron. |
| [AfxGlutWireRhombicDodecahedron](#afxglutrhombicdodecahedron) | Draws a wireframe rhombic dodecahedron. |
| [AfxGlutSolidSphere](#afxglutsphere) | Renders a solid sphere centered at the origin of the modeling coordinate system. |
| [AfxGlutWireSphere](#afxglutsphere) | Renders a wireframe sphere centered at the origin of the modeling coordinate system. |
| [AfxGlutSolidTeapot](#afxglutteapot) | Renders a solid teapot of the desired size, centered at the origin. |
| [AfxGlutWireTeapot](#afxglutteapot) | Renders a wireframe teapot of the desired size, centered at the origin. |
| [AfxGlutSolidTetrahedron](#afxgluttetrahedron) | Renders a solid tetrahedron. |
| [AfxGlutWireTetrahedron](#afxgluttetrahedron) | Renders a wireframe tetrahedron. |
| [AfxGlutSolidTorus](#afxgluttorus) | Renders a solid torus (doughnut shape). |
| [AfxGlutWireTorus](#afxgluttorus) | Renders a wireframe torus (doughnut shape). |

---

## <a name="afxglutcone"></a>AfxGlutSolidCone / AfxGlutWireCone

Renders a right circular cone with a base centered at the origin and in the X-Y plane and its tip on the positive Z-axis. The wire cone is rendered with triangular elements.

```
SUB AfxGlutSolidCone (BYVAL dbase AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```
```
SUB AfxGlutWireCone (BYVAL dbase AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dbase* | The desired radius of the base of the cone. |
| *height* | The desired height of the cone. |
| *slices* | The desired number of slices around the cone. |
| *stacks* | The desired number of segments between the base and the tip of the cone (the number of points, including the tip, is stacks + 1). |

---

## <a name="afxglutcube"></a>AfxGlutSolidCube / AfxGlutWireCube

Renders a cube of the desired size, centered at the origin. Its faces are normal to the coordinate directions.

```
SUB AfxGlutSolidCube (BYVAL dbSize AS DOUBLE)
SUB AfxGlutWireCube (BYVAL dbSize AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dbSize* | The desired length of an edge of the cube. |

---

## <a name="afxglutcylinder"></a>AfxGlutSolidCylinder / AfxGlutWireCylinder

Draws a cylinder.

```
SUB AfxGlutSolidCylinder (BYVAL radius AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```
```
SUB AfxGlutWireCylinder (BYVAL radius AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *radius* | The desired radius of the cylinder. |
| *height* | The desired height of the cylinder. |
| *slices* | The desired number of slices around the cylinder. |
| *stacks* | The desired number of segments between the base and the top of the cylinder (the number of points, including the tip, is stacks + 1). |

---

## <a name="afxglutdodecahedron"></a>AfxGlutSolidDodecahedron / AfxGlutWireDodecahedron

Renders a dodecahedron whose corners are each a distance of sqrt(3) from the origin. The length of each side is sqrt(5)-1. There are twenty corners; interestingly enough, eight of them coincide with the corners of a cube with sizes of length 2.

```
SUB AfxGlutSolidDodecahedron
SUB AfxGlutWireDodecahedron
```
---

## <a name="afxgluticosahedron"></a>AfxGlutSolidIcosahedron / AfxGlutWireIcosahedron

Renders an icosahedron whose corners are each a unit distance from the origin. The length of each side is slightly greater than one. Two of the corners lie on the positive and negative X-axes.

```
SUB AfxGlutSolidIcosahedron
SUB AfxGlutWireIcosahedron
```
---

## <a name="afxglutoctahedron"></a>AfxGlutSolidOctahedron / AfxGlutWireOctahedron

Renders an octahedron whose corners are each a distance of one from the origin. The length of each side is sqrt(2). The corners are on the positive and negative coordinate axes.

```
SUB AfxGlutSolidOctahedron
SUB AfxGlutWireOctahedron
```
---

## <a name="afxglutrhombicdodecahedron"></a>AfxGlutSolidRhombicDodecahedron / AfxGlutWireRhombicDodecahedron

Renders a rhombic dodecahedron whose corners are at most a distance of one from the origin. The rhombic dodecahedron has faces which are identical rhombuses (rhombi?) but which have some vertices at which three faces meet and some vertices at which four faces meet. The length of each side is sqrt(3)/2. Vertices at which four faces meet are found at (0, 0, +/- 1) and (+/- sqrt(2)/2, +/- sqrt(2)/2, 0). 

```
SUB AfxGlutSolidRhombicDodecahedron
SUB AfxGlutWireRhombicDodecahedron
```
---

## <a name="afxglutsphere"></a>AfxGlutSolidSphere / AfxGlutWireSphere

Renders a sphere centered at the origin of the modeling coordinate system. The north and south poles of the sphere are on the positive and negative Z-axes respectively and the prime meridian crosses the positive X-axis.

```
SUB AfxGlutSolidSphere (BYVAL radius AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)
SUB AfxGlutWireSphere (BYVAL radius AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *radius* | The desired radius of the sphere. |
| *slices* | The desired number of slices (divisions in the longitudinal direction) in the sphere. |
| *stacks* | The desired number of stacks (divisions in the latitudinal direction) in the sphere. The number of points in this direction, including the north and south poles, is stacks+1. |

---

## <a name="afxglutteapot"></a>AfxGlutSolidTeapot / AfxGlutWireTeapot

Renders a solid teapot of the desired size, centered at the origin.

| Parameter  | Description |
| ---------- | ----------- |
| *dbSize* | The desired size of the teapot. |

```
SUB AfxGlutSolidTeapot
SUB AfxGlutWireTeapot
```
---

## <a name="afxgluttetrahedron"></a>AfxGlutSolidTetrahedron / AfxGlutWireTetrahedron

Renders a tetrahedron whose corners are each a distance of one from the origin. The length of each side is 2/3 sqrt(6). One corner is on the positive X-axis and another is in the X-Y plane with a positive Y-coordinate.

| Parameter  | Description |
| ---------- | ----------- |
| *dbSize* | The desired size of the teapot. |

```
SUB AfxGlutSolidTetrahedron
SUB AfxGlutWireTetrahedron
```
---

## <a name="afxgluttorus"></a>AfxGlutSolidTorus / AfxGlutWireTorus

Renders a torus centered at the origin of the modeling coordinate system. The torus is circularly symmetric about the Z-axis and starts at the positive X-axis.

```
SUB AfxGlutSolidTorus (BYVAL innerRadius AS DOUBLE, BYVAL outerRadius AS DOUBLE, _
   BYVAL nsides AS LONG, BYVAL rings AS LONG)
```
```
SUB AfxGlutWireTorus (BYVAL innerRadius AS DOUBLE, BYVAL outerRadius AS DOUBLE, _
   BYVAL nsides AS LONG, BYVAL rings AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *innerRadius* | Inner radius of the torus. |
| *outerRadius* | Outer radius of the torus. |
| *nsides* | Number of sides for each radial section. |
| *rings* | Number of radial divisions for the torus. |

# Example

The following example draws a teapot:

```
' ########################################################################################
' Microsoft Windows
' File: GLCTX_Teapot.bas
' Contents: Graphic Control OpenGL - Teapot
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "AfxNova/CWindow.inc"
#INCLUDE ONCE "Afx/AfxGlut.inc"
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

CONST GL_WINDOWWIDTH   = 600                 ' Window width
CONST GL_WINDOWHEIGHT  = 400                 ' Window height
CONST GL_WindowCaption = "OpenGL - Teapot"   ' Window caption
CONST IDC_GRCTX = 1001

' =======================================================================================
' OpenGL class
' =======================================================================================
TYPE CtxOgl

   Private:
      m_pGraphCtx AS CGraphCtx PTR

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

   glEnable(GL_LIGHTING)
   glEnable(GL_LIGHT0)
   glEnable(GL_DEPTH_TEST)

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
   glPushMatrix
   gluLookAt(0,2.5,-7, 0,0,0, 0,1,0)
   AfxGlutSolidTeapot(2.5)
   glPopMatrix
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
      ' // creation of the window. If the application returns –1, the window is destroyed
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
```
---
