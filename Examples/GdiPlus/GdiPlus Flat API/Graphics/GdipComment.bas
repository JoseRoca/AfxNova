' ########################################################################################
' Microsoft Windows
' File: GdipComment.bas
' Contents: GDI+ Flat API - GdipComment example
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define _WIN32_WINNT &h0602
#INCLUDE ONCE "AfxNova/AfxGdiPlus.bi"
#INCLUDE ONCE "AfxNova/ARGB_Colors.bi"
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
' Adds a comment to a metafile
' ========================================================================================
SUB Example_GdipComment (BYVAL hdc AS HDC)

   DIM hStatus AS LONG

   ' // Create an IStream to hold the metafile
   DIM stream AS IStream PTR
   hStatus = CreateStreamOnHGlobal(NULL, TRUE, @stream)
   IF hStatus <> 0 OR stream = NULL THEN
      PRINT "Failed to create stream."
      EXIT SUB
   END IF

   ' // Define frame rectangle for the metafile
   DIM rcfFrame AS GpRectF = (0.0, 0.0, 400.0, 300.0)

   ' Record metafile into stream
   DIM description AS WSTRING * 64 = "Metafile with Comment"
   DIM metafile AS GpMetafile PTR
   hStatus = GdipRecordMetafileStream(stream, hdc, EmfTypeEmfPlusOnly, @rcfFrame, _
                                      MetafileFrameUnitPixel, @description, @metafile)
   IF hStatus <> 0 OR metafile = NULL THEN
      PRINT "Failed to record metafile: "; hStatus
      stream->lpvtbl->Release(stream)
      EXIT SUB
   END IF

   ' // Create graphics object for recording
   DIM graphics AS GpGraphics PTR
   hStatus = GdipGetImageGraphicsContext(metafile, @graphics)
   IF hStatus <> 0 OR graphics = NULL THEN
      PRINT "Failed to get graphics context: "; hStatus
      GdipDisposeImage(metafile)
      stream->lpvtbl->Release(stream)
      EXIT SUB
   END IF

   ' // Draw something into the metafile
   DIM pen AS GpPen PTR
   hStatus = GdipCreatePen1(ARGB_GREEN, 4.0, UnitPixel, @pen)
   hStatus = GdipDrawEllipse(graphics, pen, 50, 50, 200, 150)

   ' // Add a comment to the metafile
   DIM commentText AS STRING = "This is a custom embedded comment"
   hStatus = GdipComment(graphics, LEN(commentText), STRPTR(commentText))
   IF hStatus = StatusOk THEN
      AfxMsg "Comment added successfully."
   ELSE
      AfxMsg "Failed to add comment: hStatus: " & WSTR(hStatus)
   END IF

   ' // Cleanup
   IF pen THEN GdipDeletePen(pen)
   IF graphics THEN GdipDeleteGraphics(graphics)
   IF metafile THEN GdipDisposeImage(metafile)
   IF stream THEN stream->lpvtbl->Release(stream)

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
   pWindow.Create(NULL, "GDI+ GdipComment", @WndProc)
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
   Example_GdipComment(hdc)

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
